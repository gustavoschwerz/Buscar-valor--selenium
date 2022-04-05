#!/usr/bin/env python
# coding: utf-8

# In[3]:


#drive: https://drive.google.com/drive/folders/1KmAdo593nD8J9QBaZxPOG1yxHZua4Rtv

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time

navegador = webdriver.Chrome()


# In[4]:


navegador.get('https://google.com')
time.sleep(3)

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('preço dolar', Keys.ENTER)
    
dolar = float(navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value'))
#dolar = float(dolar)
navegador.execute_script('window.open('');') # Nova aba
navegador.switch_to.window(navegador.window_handles[1]) #trocar de janela
navegador.get('https://www.google.com/')
print('Preço do dólar U${}'.format(dolar))

#time.sleep(2)
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('preço euro', Keys.ENTER)
euro = float(navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value'))
print('Preço do euro E${}'.format(euro))
#navegador.switch_to.window(navegador.window_handles[0])

navegador.execute_script('window.open('');') # Nova aba
navegador.switch_to.window(navegador.window_handles[2]) # Troca aba
#time.sleep(1)
navegador.get('https://www.melhorcambio.com/ouro-hoje')
ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
ouro = ouro.replace(',', '.')
ouro = float(ouro)
print('Preço do ouro R${}'.format(ouro))

navegador.get('https://outlook.live.com/owa/')


# In[5]:


navegador.find_element(By.XPATH, '/html/body/header/div/aside/div/nav/ul/li[2]/a').click()
navegador.find_element(By.XPATH, '//*[@id="i0116"]').send_keys('gustavo_schwerz@hotmail.com', Keys.ENTER)


# In[6]:


navegador.find_element(By.XPATH, '//*[@id="i0118"]').send_keys('henrique9973gugu')

time.sleep(0.5)

navegador.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
navegador.find_element(By.XPATH, '//*[@id="idBtn_Back"]').click()

time.sleep(2)

navegador.find_element(By.XPATH, '//*[@id="id__8"]').click()


# In[7]:


tabela = pd.read_excel('Produtos.xlsx', 'Sheet1')

display(tabela)

#print(tabela.types)

#tabela.info()
#print('\n')
#tabela.dtypes
#cotacao_tabela_ouro = tabela[tabela['Moeda'] == 'Dólar']

tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(ouro)

tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']

tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

display(tabela)

tabela.to_excel('Produtos Atualizado.xlsx', index=False) # PARA REMOVER OS INDEX!!!

#navegador.get('https://outlook.live.com/owa/'), Keys.ENTER

#navegador.find_element('By.XPATH', '/html/body/header/div/aside/div/nav/ul/li[2]/a').click

#navegador.send_Keys('gustavo_schwerz@hotmail.com'), Keys.ENTER

#navegador.send_Keys('henrique9973gugu'), Keys.ENTER

#navegador.find_element('By.XPATH', '//*[@id="idBtn_Back"]').click

#navegador.find_element('By.XPATH', '//*[@id="id__8"]'.click

#navegador.find_element('By.XPATH', '//*[@id=":mo"]/div/div').click


#navegador.quit()


# In[9]:


time.sleep(0.5)

navegador.find_element(By.TYPE, '//*[@id="docking_InitVisiblePart_0"]/div[1]/div[1]/div[1]/div/div[3]/div/div/div/div/div/div[1]/div/div/input').send_keys('gustschwerz@gmail.com'), Keys.TAB

time.sleep(0.5)

navegador.find_elements(By.XPATH, '//*[@id="TextField401"]').send_keys('Dólar, Euro e Ouro'), Keys.TAB

    texto = """Segue cotação das moedas solicitadas.

O valor do dólar é: U${:,.2f}
O valor do euro é: E${:,.2f}
O valor do outo é: R${:,.2f}

A planilha foi atualizada com sucesso.

Att. Robô
""".format(dolar, euro, ouro)

navegador.find_element(By.XPATH, '//*[@id="virtualEditScroller423"]/div/div').send_keys(texto)
                      
navegador.find_element(By.XPATH, '//*[@id="docking_InitVisiblePart_0"]/div[1]/div[3]/div[2]/div[1]/div/div/span/button[1]/span/i').click()
#navegador.quit()


# In[ ]:




