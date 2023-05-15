#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import PySimpleGUI as sg


# In[2]:


pip install PySimpleGUI


# In[4]:


excel_file="/Users/dhwanibhavankar/Library/Containers/com.microsoft.Excel/Data/Downloads/python_dashboard.xlsx"


# In[5]:


df=pd.read_excel('/Users/dhwanibhavankar/Library/Containers/com.microsoft.Excel/Data/Downloads/python_dashboard.xlsx',sheet_name='Sheet1')


# In[6]:


df


# In[7]:


sg.theme('DarkTeal9')


# In[8]:


layout = [
    [sg.Text('Please fill out the following fields:')], 
    [sg.Text('LoanID', size=(15, 1)),sg.InputText(key='Loan_ID')],
    [sg.Text('Gender', size=(15, 1)),sg.InputText(key='Gender')],
    [sg.Text('Married', size=(15, 1)),sg.Checkbox('TRUE',key='Married'),sg.Checkbox('FALSE',key='Married')],
    [sg.Text('Dependents', size=(15, 1)),sg.InputText(key='Dependents')],
    [sg.Text('Education', size=(15, 1)),sg.InputText(key='Education')],
    [sg.Text('Self_Employed', size=(15, 1)),sg.Checkbox('TRUE',key='Self_Employed'),sg.Checkbox('FALSE',key='Self_Employed')],
    [sg.Text('ApplicantIncome', size=(15, 1)),sg.InputText(key='ApplicantIncome')],
    [sg.Text('CoapplicantIncome', size=(15, 1)),sg.InputText(key='CoapplicantIncome')],
    [sg.Text('LoanAmount', size=(15, 1)),sg.InputText(key='LoanAmount')],
    [sg.Text('Loan_Amount_Term', size=(15, 1)),sg.InputText(key='Loan_Amount_Term')],
    [sg.Text('Credit_History', size=(15, 1)),sg.InputText(key='Credit_History')],
    [sg.Text('Property_Area(Urban/Semiurban/Rural)', size=(15, 1)),sg.InputText(key='Property_Area')],
    [sg.Submit(),sg. Exit ()]
]


# In[9]:


window=sg.Window("Data Entry Form",layout)


# In[ ]:


while True:
    event,values=window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        df=df.append(values,ignore_index=True)
        df.to_excel(excel_file,sheet_name='Sheet1',index=False)
        sg.popup('Data Saved !!!')
window.close()


# In[ ]:




