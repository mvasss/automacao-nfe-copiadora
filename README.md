
# 🧾 Automação de Notas Fiscais – Prefeitura de Osasco

Automação completa para **emissão** e **verificação** de Notas Fiscais de Serviço Eletrônicas (NFS-e) no portal da prefeitura de Osasco, utilizando Python + Selenium.

Ideal para empresas que emitem ou conferem notas com frequência, como copiadoras, prestadores de serviço, contabilidades, etc.

---

## ⚙️ Funcionalidades

✔️ Emissão automática de NFS-e com base em planilha Excel  
✔️ Verificação de notas já emitidas para CPF/CNPJ com conferência de valor  
✔️ Tratamento de múltiplos resultados, pop-ups e erros comuns do site  
✔️ Registro automático do status ("sim"/"não") na planilha de entrada

---

## 🛠️ Tecnologias utilizadas

- Python 3.x  
- Selenium  
- Pandas  
- ChromeDriver

---

## 📝 Pré-requisitos

- Python instalado  
- Google Chrome instalado  
- ChromeDriver compatível com a versão do seu navegador  
- Instalar dependências com:

```bash
pip install -r requirements.txt
```

---

## 📂 Estrutura dos scripts

- `emissor_nfe_selenium.py` → Faz a **emissão** das notas com base no Excel  
- `pesquisa_nfe.py` → Faz a **verificação** se a nota foi emitida corretamente  
    - Lê nome, CPF/CNPJ e valor  
    - Consulta no site da prefeitura  
    - Escreve na planilha se a nota existe ou não

---

## 📄 Exemplo de planilha de entrada (Excel)

| Nome do Pagador | CPF/CNPJ       | Valor |
|-----------------|----------------|-------|
| João da Silva   | 12345678900    | 150.00|
| Maria Oliveira  | 98765432100    | 50.00 |

> A planilha deve estar no mesmo diretório ou com o caminho especificado no script.

---

## ▶️ Como usar

### Emissão:
1. Preencha a planilha com os dados dos clientes.  
2. Edite o script com suas credenciais de acesso.  
3. Rode o script:

```bash
python emissor_nfe_selenium.py
```

### Verificação:
1. Use a mesma planilha (ou uma nova com os dados a verificar).  
2. Rode o script:

```bash
python pesquisa_nfe.py
```

---

## 🧠 Observações

- Os scripts são feitos sob medida para o portal da **Prefeitura de Osasco (SP)**  
- Caso o portal mude, os seletores podem precisar de ajustes  
- O navegador será controlado automaticamente – não mexa durante a execução

---

## 📫 Contribuições e dúvidas

Sinta-se à vontade para abrir issues ou fazer pull requests.  
Dúvidas, sugestões ou melhorias são sempre bem-vindas!
