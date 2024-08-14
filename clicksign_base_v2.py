#!/usr/bin/env python
# coding: utf-8

# In[4]:


from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time as time


servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)


#Pagina pra onde sera redirecionado
navegador.maximize_window()
navegador.get("https://app.clicksign.com")



# In[5]:


#pega o campo email e digita
navegador.find_element('xpath','//*[@data-testid="emailField"]' ).send_keys("atendimentovidasaude01@gmail.com")


# In[ ]:


navegador.implicitly_wait(10)


# In[ ]:


#pega o campo senha e digita
navegador.find_element('xpath','//*[@data-testid="passwordField"]').send_keys("RMvs2024@")


# In[ ]:


navegador.implicitly_wait(10)


# In[ ]:


navegador.find_element('xpath','//*[@id="login_email_form"]/button').click()


# In[ ]:


time.sleep(10)


# In[ ]:


#Pagina pra onde sera redirecionado
navegador.get("https://app.clicksign.com/accounts/565988/dashboard")


# In[ ]:


navegador.implicitly_wait(25)


# In[ ]:


navegador.find_element('xpath','//*[@class="_XButton_1p73z_1 _modelDefault_1p73z_153 _modelDefaultDesign_1p73z_161 _modelDefaultDesignDefault_1p73z_161 _radiusRounded_1p73z_78 _isFullWidth_1p73z_43"]').click()


# In[ ]:


time.sleep(30)


# In[ ]:


#APOS USUARIO CARREGAR ARQUIVO
navegador.find_element('xpath','//*[@data-testid="nextStepBtn"]').click()


# In[ ]:


#ADICIONAR SIGNATARIO
navegador.find_element('xpath','//*[@data-testid="addSignerButton"]').click()


# In[ ]:


time.sleep(70)


# In[ ]:


#PEGA OS DADOS DA PLANILHA E ADICIONA NO PROGRAMA
from docx import Document
from datetime import datetime 
import pandas as pd

tabela = pd.read_excel("Informações.xlsx", sheet_name='Planilha1')

for linha in tabela.index:

    nomeCliente = tabela.loc[linha, "Nome"]
    dataNascimento = str(tabela.loc[linha, "DataNascimento"].strftime( '%d/%m/%Y' ))
    cpfCliente = str(tabela.loc[linha, "CPF"])
    emailCliente = str(tabela.loc[linha, "Email"])


# In[ ]:


#CLICA NO BOTÃO PROXIMO
navegador.find_element('xpath','//*[@data-testid="nextStepAddSignerModalButton"]').click()


# In[ ]:


time.sleep(1)


# In[ ]:


#SIGNATARIO DEVE ASSINAR - SELEÇÃO
navegador.find_element('xpath','//*[@data-testid="signerSignAsSelectPlaceholder"]' ).click()


# In[ ]:


time.sleep(1)


# In[ ]:


#ASSINAR - CHECKBOX
navegador.find_element('xpath','//*[@data-testid="selectFieldSignOptionCheckbox"]/div' ).click()


# In[ ]:


time.sleep(1)


# In[ ]:


#BOTÃO AVANÇAR
navegador.find_element('xpath','//*[@data-testid="nextStepAddSignerModalButton"]' ).click()


# In[ ]:


time.sleep(6)


# In[ ]:


#BOTÃO AVANÇAR
navegador.find_element('xpath','//*[@data-testid="nextStepBtn"]' ).click()


# In[ ]:


time.sleep(2)


# In[ ]:


#BOTÃO AVANÇAR
navegador.find_element('xpath','//*[@data-testid="nextStepBtn"]' ).click()


# In[ ]:


time.sleep(10)


# In[ ]:


#pega o campo DATA DE NASCIMENTO e digita
navegador.find_element('xpath','//*[@data-testid="documentMessageField"]' ).send_keys("Olá "+nomeCliente+", tudo bem? Segue em anexo o Termo de Adesão referente ao Plano Aderido com a DG Vida Saúde. ")
navegador.find_element('xpath','//*[@data-testid="documentMessageField"]' ).send_keys("Preciso que assine o documento, clicando no botão abaixo. DG Vida Saúde agradece. Estamos a disposição para eventuais dúvidas que venha surgir!")


# In[ ]:


time.sleep(2)


# In[ ]:


#BOTÃO ENVIAR DOCUMENTO
navegador.find_element('xpath','//*[@id="sendMessage"]' ).click()


# In[ ]:


time.sleep(10)


# In[ ]:


navegador.close()

