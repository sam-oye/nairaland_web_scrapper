#!/usr/bin/env python
# coding: utf-8

# In[145]:


import requests
from bs4 import BeautifulSoup
import os
import win32com.client as win32
import datetime



# In[146]:


r = requests.get("https://www.nairaland.com/")
c=r.content


# In[147]:


soup=BeautifulSoup(c, "html.parser")


# In[148]:


all=soup.find_all("td",{"class":"featured w"})


# In[ ]:





# In[149]:


pages=all[0].find_all("a")
k2=[]
for i3 in range(65):
    k2.append(all[0].find_all("a")[i3].text)
print(len(pages))


# In[150]:


#print(pages[0])


# In[ ]:


k1=[]
k4=[]
k5=[]
for i in range(len(pages)-10):
    
    j=pages[i]["href"]
    r1 = requests.get(j)
    c1=r1.content
    soup1=BeautifulSoup(c1, "html.parser")
    all1=soup1.find_all("p",{"class":"bold"})
    all3=soup1.find_all("a")[9].text
    k5.append(all3)
    k=[]
    for j in all1[0]:
        k.append(j)
        #print(j)
    try:
        k1.append(k[7])
        k4.append(k[4])
    except IndexError:
        
        k1.append(k[5])
        k4.append(k[2])
    
    #pages1=all1[0].find_all("p")
    #print(pages1)
#for l in k1:


# In[ ]:


k3=[]
for i1 in range(len(k1)):
    k3.append([k2[i1],k1[i1],k5[i1]])

    


# In[ ]:



j=pages[60]["href"]
r2 = requests.get(j)
c2=r2.content
soup1=BeautifulSoup(c2, "html.parser")
#ll2=soup1.find_all("p",{"class":"bold"})
#print(all2)

all3=soup1.find_all("a")[9].text
print(all3)


# In[ ]:


xl = win32.Dispatch('Excel.Application')
xl.Visible = True

wb = xl.Workbooks.Add()
wb.Sheets.Add().Name="With Project"
ws_sheet1 = wb.Worksheets('With Project') 


# In[ ]:


ws_sheet1.Cells(1,"A").Value = "TOPICS"
ws_sheet1.Cells(1,"B").Value = "VIEWS"
ws_sheet1.Cells(1,"C").Value = "SECTIONS"


# In[ ]:



for i in range(len(k3)-2):
    
    i=i+2
    
    ws_sheet1.Cells(i,"A").Value = k3[i][0]
    ws_sheet1.Cells(i,"B").Value = k3[i][1]
    ws_sheet1.Cells(i,"C").Value = k3[i][2]
    
    
x = datetime.datetime.now()  
x=str(x)
x=x.replace("-", "")
x=x.replace(" ", "")
x=x.replace(":", "")
x=x.replace(".", "")


# In[ ]:


wb.SaveAs(Filename=os.path.join(os.getcwd(), str(x)+ ".xlsx"))


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




