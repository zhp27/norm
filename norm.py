# -*- coding: utf-8 -*-
"""
Created on Tue Mar 16 18:21:59 2021

@author: raha_
"""
#read docx file and break it to pragraphs
from docx import Document
import re
document = Document('before_1.docx')
par=document.paragraphs


#text canonization
def cleanText(text):
    puncts='_-()"”“©—'
    for sym in puncts:
        text= text.replace(sym,' ')
    reps='%&'
    for sym in reps:
        if sym=='%':
            text= text.replace(sym,' percent ')
        if sym=='&':
            text= text.replace(sym,' and ')
        if sym=='/':
            text= text.replace(sym,' or ')
    return text
def CountryAbv(text):    
    abbrevs={'USA':'United States','GB':'Great Britain'}
    for abbrev in abbrevs:
        text= text.replace(abbrev,abbrevs[abbrev])
    return text

def getText(doc):   
    fullText = []
    for para in doc:
        fullText.append(para.text)
    return '\n'.join(fullText)  
def curRemove(text):
    t=''
    sp=text.split()
    puncts='€$,.;'
    for w in sp:        
        if w.startswith('$'):
            w+=' Dollar '
            for p in puncts:
                w=w.replace(p, "")                         
            
        t=t +' ' +w
        if w.startswith('€'):
            w+=' Euro '
            for p in puncts:
                w=w.replace(p, "") 
            print(w)               
           
        t=t +' ' +w        
    return t
    
    
T=getText(par)  
CurP=T.lower()
CurPN=cleanText(CurP)
CurPNC=CountryAbv(CurPN)
curRemove(CurPNC)
    
    
    
    
    
    

