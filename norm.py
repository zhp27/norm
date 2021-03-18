# -*- coding: utf-8 -*-
"""
Created on Wed Mar 17 11:23:32 2021

@author: raha_
"""
#text canonization
from docx import Document
import re
import spacy
from nltk import sent_tokenize
import inflect
n2s = inflect.engine()



#perform NER
def NER_d(sentence):    
    doc = nlp(sentence) 
    ne=[]
    for ent in doc.ents:
        
        ne.append(ent)
       # print(ent.text, ent.start_char, ent.end_char, ent.label_)
        
    return ne

#convert times to string
def convert_num(text):
    timer=re.match(r'^(([0-1]{0,1}[0-9]( )?(AM|am|aM|Am|PM|pm|pM|Pm))|(([0]?[1-9]|1[0-2])(:|\.)[0-5][0-9]( )?(AM|am|aM|Am|PM|pm|pM|Pm))|(([0]?[0-9]|1[0-9]|2[0-3])(:|\.)[0-5][0-9]))', text)    
    if timer:
        
        time=timer.group()
        time1=time.replace(':', ' ')
        
        text= text.replace(time,time1)
    
    mo=re.findall(r'\d+', text)
    if mo:        
        for item in mo:
            sa=n2s.number_to_words(item)
            text= text.replace(item,sa)
    return text
        

    

#removing symbolls
def clean_text(text):
    text = text.replace('\r', ' ')
    text = text.replace('\n', ' ')
    puncts='_-()"”“©—*:,.'
    for sym in puncts:
        text= text.replace(sym,' ')
    reps='%&/#+'
    for sym in reps:
        if sym=='%':
            text= text.replace(sym,' percent ')
        if sym=='&':
            text= text.replace(sym,' and ')
        if sym=='/':
            text= text.replace(sym,' or ')
        if sym=='#':
            text= text.replace(sym,' number ')
        if sym=='+':
            text= text.replace(sym,' plus ')
    return text
#Remove country abriviations
def Country_abv(text):    
    abbrevs={'USA':'United States','GB':'Great Britain'}
    for abbrev in abbrevs:
        text= text.replace(abbrev,abbrevs[abbrev])
    return text

#convert money symbols
def get_text(doc):   
    fullText = []
    for para in doc:
        fullText.append(para.text)
    return '\n'.join(fullText)  
def cur_remove(text):
    t=''
    sp=text.split()
    puncts='€$,.;'
    for w in sp: 
        #print(w)
        if w.startswith('$'):
           # print(w)
            w+=' Dollar '            
            for p in puncts:
                if p=='$':
                    w=w.replace(p, "")
                    #print(w)                  
            
        
        if w.startswith('€'):
            w+=' Euro '
            for p in puncts:
                if p=='€':
                    w=w.replace(p, "")                          
         
        t=t +' ' +w        
    return t

'''def remove_th(text):
    
    text.replace(r'/(\d+)(st|nd|rd|th)/', r'1');
    print(text)'''
    
#concatenate lists and bullet points and removing empty line
def file_creation(doc):    
    Paras = normDoc.paragraphs   
    l=len(Paras) 
    fnormDoc= Document()
    p1=normDoc.paragraphs[0] 
    #print(l)
    st=str(p1.text)
       #print(st)
    fnormDoc.add_paragraph(st)
    
    for i in range(1,l):
       p1=normDoc.paragraphs[i] 
       l2=len(p1.text)
       #join small lines of lists 
       if l2<50:
           continue
       for j in  range(i+1,l):
           p=normDoc.paragraphs[j] 
           l1=len(p.text)         
           if l1>50: 
               break          
               
           else:
               p1.add_run(p.text)
       
       st=str(p1.text)
       #print(st)
       fnormDoc.add_paragraph(st)
       fnormDoc.save("norm.docx")

    
#main code

#read the doc
doc = Document('test3.docx')
Paras = doc.paragraphs
normDoc= Document()

for p in Paras: 
    #integrate pargraphs
    p.keep_with_next=True 
    p.keep_together=True 
    
    for run in p.runs:
        
        T=run.text  
              
        CurPNCC=cur_remove(T)
        Cur=convert_num(CurPNCC)
        #CurP=Cur.lower()  
             
        CurPN=clean_text(Cur)
        CurPNC=Country_abv(CurPN) 
                
        normDoc.add_paragraph(CurPNC)
file_creation(normDoc)
    
    #normDoc.save("a.docx")
    
        
       
  
        
        
        


