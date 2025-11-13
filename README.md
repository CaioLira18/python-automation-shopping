# 📊 Sistema de Relatório de Vendas por Loja

Sistema automatizado em Python para análise de vendas e envio de relatórios por email usando dados do Excel.

## 📋 Descrição

Este projeto analisa dados de vendas de múltiplas lojas, calcula métricas importantes (faturamento, quantidade vendida e ticket médio) e envia automaticamente um relatório formatado por email via Outlook.

## ✨ Funcionalidades

- 📈 **Análise de Faturamento**: Calcula o valor total de vendas por loja
- 📦 **Controle de Quantidade**: Soma produtos vendidos por estabelecimento
- 💰 **Ticket Médio**: Calcula o valor médio por produto em cada loja
- 📧 **Envio Automático**: Dispara relatório formatado em HTML por email

## 🔧 Requisitos

### Bibliotecas Python
```bash
pip install pandas openpyxl pywin32
```

### Requisitos do Sistema
- Python 3.7 ou superior
- Microsoft Outlook instalado e configurado
- Windows (necessário para integração com Outlook)

## 📁 Estrutura de Arquivos

```
projeto/
│
├── script.py           # Script principal
└── Vendas.xlsx         # Base de dados (deve conter as colunas necessárias)
```

### Formato do arquivo `Vendas.xlsx`

O arquivo Excel deve conter as seguintes colunas:
- `ID Loja`: Identificador da loja
- `Valor Final`: Valor total da venda
- `Quantidade`: Quantidade de produtos vendidos

## 🚀 Como Usar

1. **Prepare o arquivo de dados**
   - Certifique-se de que o arquivo `Vendas.xlsx` está no mesmo diretório do script
   - Verifique se as colunas estão nomeadas corretamente

2. **Configure o destinatário**
   ```python
   mail.To = 'seu-email@exemplo.com'  # Altere para o email desejado
   ```

3. **Execute o script**
   ```bash
   python script.py
   ```

4. **Verifique a saída**
   - Os resultados serão exibidos no console
   - Um email será enviado automaticamente via Outlook

## 📊 Exemplo de Saída

```
ID Loja  Valor Final
1        15000.00
2        23500.00
3        18750.00
--------------------------------------------------
ID Loja  Quantidade
1        250
2        380
3        310
--------------------------------------------------
        Valor Final
ID Loja            
1         60.00
2         61.84
3         60.48
--------------------------------------------------
Email Enviado.
```

## ⚙️ Personalização

### Alterar o assunto do email
```python
mail.Subject = 'Seu novo assunto aqui'
```

### Modificar o template do email
Edite a variável `mail.HTMLBody` com seu próprio HTML:
```python
mail.HTMLBody = f'''
<p>Seu texto personalizado</p>
{faturamento.to_html()}
'''
```

### Adicionar anexos
```python
mail.Attachments.Add('caminho/para/arquivo.pdf')
```

## 🐛 Solução de Problemas

**Erro: "Arquivo não encontrado"**
- Verifique se `Vendas.xlsx` está no diretório correto

**Erro ao enviar email**
- Confirme que o Outlook está instalado e configurado
- Execute o script com permissões de administrador se necessário

**Erro: "Coluna não encontrada"**
- Verifique os nomes exatos das colunas no arquivo Excel

## 📝 Notas

- O script usa F-strings implícitas no HTML. Para funcionar corretamente, adicione `f` antes das aspas triplas:
  ```python
  mail.HTMLBody = f'''...'''
  ```
- O Outlook pode solicitar permissão na primeira execução
- Certifique-se de ter uma conta configurada no Outlook

## 📄 Licença

Este projeto é de código aberto e está disponível para uso livre.

---

**Desenvolvido para automação de relatórios comerciais** 🚀
