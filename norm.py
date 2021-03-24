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

from docx.enum.style import WD_STYLE_TYPE

#replace the inflict
from num2words import num2words

#perform NER 
def NER_d(sentence):    
    doc = nlp(sentence) 
    ne=[]
    for ent in doc.ents:
        
        ne.append(ent)
       # print(ent.text, ent.start_char, ent.end_char, ent.label_)
        
    return ne

#convert numbers and dates to strings using regex
def convert_num(text):
    timer=re.match(r'^(([0-1]{0,1}[0-9]( )?(AM|am|aM|Am|PM|pm|pM|Pm))|(([0]?[1-9]|1[0-2])(:|\.)[0-5][0-9]( )?(AM|am|aM|Am|PM|pm|pM|Pm))|(([0]?[0-9]|1[0-9]|2[0-3])(:|\.)[0-5][0-9]))', text)    
    if timer:        
        time=timer.group()
        time1=time.replace(':', ' ')
        
        text= text.replace(time,time1)
    
    mo=re.findall(r'\d+', text)
    if mo:        
        for item in mo:
            sa=num2words(item)            
            sa=sa.replace('\r',' ')
            text= text.replace(item,sa)       
            
             
    return text
        

 

#removing symbolls and runs
def clean_text(text):
    text = text.replace('\r', ' ')
    text = text.replace('\n', ' ')
    puncts='_-()"”“©—*:'
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
        if sym=='=':
            text= text.replace(sym,' equals ')
    return text
#Remove country abriviations
def Country_abv(text):    
    acronyms = [
        ('U.S.', 'United States'), 
        ('USA', 'United States'),
        ('US', 'United States'),
        ('U.S.', 'United States'),
        ('UK', 'United Kingdom'),
        ('U.K.', 'United Kingdom'),
        ('Great Britain', 'United Kingdom'),
        ('Britain', 'United Kingdom'),
        
        
    ]
    pattern = '|'.join('(%s)' % re.escape(match) for match, replacement in acronyms)
    substitutions = [match for replacement, match in acronyms]
    replace = lambda m: substitutions[m.lastindex - 1]
    return re.sub(pattern, replace, text)
def lang_abv(text):    
    acronyms = [
        ('i.e.', 'for example'),
        ('e.g.', 'for instance') ,      
        ('’ve',' have' ),
        ('’re', ' are'),        
        ('I’m',' I am' ),
        ('n’t', ' not'),
        ('’s', ' is')
        
    ]
    pattern = '|'.join('(%s)' % re.escape(match) for match, replacement in acronyms)
    substitutions = [match for replacement, match in acronyms]
    replace = lambda m: substitutions[m.lastindex - 1]
    return re.sub(pattern, replace, text)

#convert money symbols
def get_text(doc):   
    fullText = []
    for para in doc:
        fullText.append(para.text)
    return '\n'.join(fullText)  
#replace currency values such as $ with their respective text
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
    #connct succesive lines
    i=0
    while i<l-1:
       i=i+1    
       #print(i)
       p1=normDoc.paragraphs[i] 
       l2=len(p1.text)       
       if l2<50:           
           for j in  range(i+1,l):
               p=normDoc.paragraphs[j] 
               l1=len(p.text)         
               if l1>=50: 
                   i=j
                   p1.add_run(p.text)
                   
                   break          
                   
               else:
                   p1.add_run(p.text) 
                   p1.add_run(',')
                   i=j
                   
       else:
           p1=normDoc.paragraphs[i]
       #print(p1.text)
       st=str(p1.text)
       #print(st)
       fnormDoc.add_paragraph(st)
       fnormDoc.save("bench3_n.docx")

def list_norm(pars):   
    noList= Document()

    for p in Paras:
        for run in p.runs:
            x=re.sub(r'\w[.)]\s*', '', run.text) 
            noList.add_paragraph(x)
    return noList
#removing all the styles
def remove_styles(doc):
        
    st=doc.styles
    for s in st:    
        strs=s.name        
        st[strs].delete
    return doc

#main code

doc = Document('bench3_un.docx')
Paras = doc.paragraphs
docwl = list_norm(Paras)
doc =remove_styles(docwl)

normDoc= Document()

for p in Paras:
    p.keep_together=True 
    p.keep_with_next=True
    for run in p.runs:
        run.text.replace('\n\n', '\n')  
        run.text.replace('\r\r', '\r') 
        T=run.text       
        CurPNCC=cur_remove(T)
        Cur=convert_num(CurPNCC)
        #CurP=Cur.lower()   
        CurPNC=Country_abv(Cur) 
        CurL=lang_abv(CurPNC)
            
        CurPN=clean_text(CurL)
        
                
        normDoc.add_paragraph(CurPN)
file_creation(normDoc)

       


    
    
       
    
        
        
       
  
        
        
        


