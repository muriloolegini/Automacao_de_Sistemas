# Automação de Sistemas e Processos com Python

### Desafio:

Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior

E-mail da diretoria: seugmail+diretoria@gmail.com<br>
Local onde o sistema disponibiliza as vendas do dia anterior: https://exemplolink

Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

``` bash

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 0.9

# Passo 1: Entrar no sistema da empresa (no caso é o link do drive)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://exemplolink")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(3.5)

# Passo 2: Navegar no sistema e encontrar a base de vendas (entrar na pasta exportar)
pyautogui.click(x=430, y=498, clicks=2)
time.sleep(1)

# Passo 3: Fazer o download da base de vendas
pyautogui.click(button='right'); pyautogui.mouseDown(x=477, y=663); pyautogui.click() # clicar no arquivo com o botão direito do mouse e clicar em Download.
time.sleep(3.5) # esperar o download acabar
pyautogui.hotkey("ctrl", "w") # Fechar a janela

```

---

### Vamos agora ler o arquivo baixado para pegar os indicadores

- Faturamento
- Quantidade de Produtos

``` bash

# Passo 4: Importar a base de vendas pro Python
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Downloads\Vendas - Dez.xlsx")
display(tabela)

```
![image](https://user-images.githubusercontent.com/22229127/194772620-a5c5fbdd-8b5e-4e6d-b3bc-79602db397dc.png)

``` bash

# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)

```

---

### Vamos agora enviar um e-mail

``` bash

# Passo 6: Enviar um e-mail para a diretoria com os indicadores de venda
# 
# abrir aba
pyautogui.hotkey("ctrl", "t")

# entrar no link do email - https://mail.google.com/mail/u/0/#inbox
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# clicar no botão escrever
pyautogui.click(x=146, y=222)
time.sleep(0.5)

# preencher as informações do e-mail

pyautogui.write("seuemail@gmail.com")
pyautogui.press("tab") # selecionar o email
time.sleep(0.5)
pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab") # pular para o campo de corpo do email

texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Dr Python
"""

# formatação dos números (moeda, dinheiro)

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")
time.sleep(5)
pyautogui.hotkey("ctrl", "w") # Fechar a janela

```

---

[⬆ Voltar ao topo](https://github.com/muriloolegini/Automacao_de_Sistemas)
