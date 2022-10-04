#!/usr/bin/env python
# coding: utf-8

# In[87]:


import pyautogui
import pyperclip
import time
import pandas as pd


pyautogui.PAUSE = 4

pyautogui.press('winleft')
    
pyautogui.write('firefox')
pyautogui.press('enter')

link = 'https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym'
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')

pyautogui.press('enter')
time.sleep(15)

pyautogui.click(x=1100, y=361)    
pyautogui.click(x=1162, y=165)
pyautogui.click(x=941, y=578)
time.sleep(15)

df = pd.read_excel(r'C:\Users\marqu\Downloads\Vendas - Dez.xlsx')
display(df)

faturamento = df['Valor Final'].sum()
qdt_produtos = df['Quantidade'].sum()

display(faturamento,qdt_produtos)


# 

# pyautogui.PAUSE = 6
# 
# pyautogui.position()

# In[88]:


pyautogui.hotkey('ctrl','t')
pyautogui.write(r'mail.google.com')
pyautogui.press('enter')
time.sleep(10)

pyautogui.click(x=103, y=175)
pyautogui.write('pythonimpressionador@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = 'Relatório de Vendas de ontem'
pyperclip.copy(assunto)
pyautogui.hotkey('crtl','v')

texto = f'''
    Prezados, bom dia
    
    O faturamento de ontem foi de R$ {faturamento:,.2f}
    A quantidade de produtos foi {qdt_produtos:,}
    
    abc
    
    Marcos Barbosa
'''
pyperclip.copy(texto)
pyautogui.hotkey('crtl', 'v')
pyautogui.hotkey('crtl','enter')


# 

# 

# 

# 
